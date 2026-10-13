"""Stable action outcomes shared by MCP actions and read-only wait tools."""

from __future__ import annotations

import json
import re


class ActionError(RuntimeError):
    """A machine-readable failure, including through legacy string handlers."""

    def __init__(self, code: str, message: str, retryable: bool = False):
        self.code = code
        self.retryable = retryable
        self.message = message
        super().__init__(f"[{code}] {message}")


def result(ok: bool, code: str, message: str, *, verified: bool = False,
           retryable: bool = False, data: dict | None = None) -> dict:
    return {
        "ok": ok, "code": code, "message": message,
        "verified": verified, "retryable": retryable, "data": data or {},
    }


ERROR_CODES = {
    "control_not_found", "ambiguous_match", "window_invalid", "action_failed",
    "result_unconfirmed", "invalid_argument", "operation_blocked", "control_disabled",
}


def _failure_code(message: str) -> str | None:
    # Typed markers survive the Controller's legacy exception-to-string conversion.
    marker = re.search(r"\[([a-z_]+)\]", message)
    if marker and marker.group(1) in ERROR_CODES:
        return marker.group(1)
    lowered = message.lower()
    if "mode is learn" in lowered or "requires explicit confirmation" in lowered:
        return "operation_blocked"
    if any(value in lowered for value in (
        "could not be read back", "result unconfirmed", "not verified",
    )):
        return "result_unconfirmed"
    if "ambiguous" in lowered or "not unique" in lowered or "不唯一" in message:
        return "ambiguous_match"
    if "window" in lowered and any(value in lowered for value in (
        "not found", "not visible", "could not be activated", "no window",
    )):
        return "window_invalid"
    if "disabled" in lowered or "控件不可点击" in message:
        return "control_disabled"
    if any(value in lowered for value in (
        "control not found", "no control found", "was not found on screen",
        "ocr text was not found", "did not find the requested control",
    )) or "未找到" in message:
        return "control_not_found"
    if lowered.startswith("error") or "异常" in message or "失败" in message:
        return "action_failed"
    return None


def normalize_result(value, *, observation: bool = False) -> dict:
    """Normalize legacy returns; execution success alone is never verification.

    Preserve original JSON fields for existing consumers, including partial batch
    results. Plain-text responses move to ``message``. No action is retried here.
    """
    if isinstance(value, ActionError):
        return result(False, value.code, value.message, retryable=value.retryable)
    if isinstance(value, Exception):
        code = "invalid_argument" if isinstance(value, (ValueError, TypeError)) else (_failure_code(str(value)) or "action_failed")
        return result(False, code, str(value))
    original = value
    if isinstance(value, str):
        try:
            decoded = json.loads(value)
        except (ValueError, TypeError):
            pass
        else:
            if isinstance(decoded, dict):
                value = decoded
    if isinstance(value, dict):
        if all(key in value for key in ("ok", "code", "verified", "retryable", "data")):
            return value
        ok = value.get("ok", value.get("found", True)) is not False
        evidence = value.get("action_verification") or {}
        verified = ok and (observation or evidence.get("verified") is True or value.get("verified") is True)
        if not ok:
            default_message = "No matching control was found" if value.get("found") is False else "Action failed"
        elif verified:
            default_message = "Requested observation or readback was verified"
        else:
            default_message = "Action executed; verify the expected effect"
        message = str(value.get("message") or value.get("error")
                      or (value.get("failed_step") or {}).get("error") or default_message)
        code = "verified" if verified else "action_executed"
        if not ok:
            code = (value.get("code") if value.get("code") in ERROR_CODES else None) or _failure_code(message) or (
                "control_not_found" if value.get("found") is False else "action_failed"
            )
        envelope = result(ok, code, message, verified=verified, data=value)
        return {**value, **envelope}
    message = str(original) if original is not None else "Action executed; verify the expected effect"
    legacy_success = re.search(r"执行动作 [^\n]+ 成功[！!]?$", message) is not None
    code = None if legacy_success else _failure_code(message)
    if original is False:
        code = "action_failed"
    if code:
        # Legacy strings cannot establish whether a partial side effect occurred.
        return result(False, code, message)
    return result(True, "verified" if observation else "action_executed", message,
                  verified=observation)
