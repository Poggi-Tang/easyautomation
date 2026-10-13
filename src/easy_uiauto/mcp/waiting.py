"""Read-only UI condition polling with injectable platform probes and clock.

The probe must perform one observation without activating a window, mutating the
UI, or starting its own retry loop. A synchronous probe cannot be interrupted;
timeouts bound polling and sleeping, and an over-budget observation is reported
as unconfirmed. A zero timeout explicitly requests a single immediate check.
"""

from __future__ import annotations

import math
import time
from collections.abc import Callable, Mapping
from copy import deepcopy

from .results import ActionError, result

Observation = dict
Probe = Callable[[], Observation]
WINDOW_FIELDS = frozenset(
    {"handle", "title", "class_name", "process_id", "bounds", "enabled", "visible"}
)


def _invalid(message: str) -> dict:
    return result(False, "invalid_argument", message)


def _poll(
    observe: Probe,
    predicate: Callable[[Observation], bool],
    *,
    condition: dict,
    timeout_seconds: float,
    interval_seconds: float,
    clock: Callable[[], float],
    sleep: Callable[[float], None],
) -> dict:
    try:
        timeout = float(timeout_seconds)
        interval = float(interval_seconds)
    except (TypeError, ValueError, OverflowError):
        return _invalid("timeout_seconds and interval_seconds must be finite numbers")
    if not math.isfinite(timeout) or not 0 <= timeout <= 120:
        return _invalid("timeout_seconds must be between 0 and 120 seconds")
    if not math.isfinite(interval) or interval <= 0:
        return _invalid("interval_seconds must be a positive finite number")

    started = clock()
    deadline = started + timeout
    attempts = 0
    observation: dict = {}
    last_code = "result_unconfirmed"

    def evidence() -> dict:
        return {
            "condition": condition,
            "observation": observation,
            "attempts": attempts,
            "elapsed_seconds": round(max(0.0, clock() - started), 6),
            "last_code": last_code,
        }

    while True:
        attempts += 1
        matched = False
        try:
            try:
                sample = observe()
            except ActionError as error:
                # Absence is an observation only when explicitly identified.
                # A stale parent window or a failed probe cannot prove absence.
                if error.code != "control_not_found":
                    raise
                sample = {"exists": False}
            if not isinstance(sample, Mapping) or not isinstance(sample.get("exists"), bool):
                raise ValueError("UI probe must return an observation with a boolean exists field")
            observation = deepcopy(dict(sample))
            last_code = "ok" if observation["exists"] else "control_not_found"
            matched = predicate(observation)
        except ActionError as error:
            return result(
                False,
                error.code,
                str(error),
                retryable=getattr(error, "retryable", False),
                data=evidence(),
            )
        except Exception as error:
            return result(
                False,
                "action_failed",
                f"Could not observe the requested UI condition: {error}",
                data=evidence(),
            )

        now = clock()
        if matched and (timeout == 0 or now <= deadline):
            return result(
                True,
                "verified",
                "Requested UI condition was observed",
                verified=True,
                data=evidence(),
            )
        remaining = deadline - now
        if timeout == 0 or remaining <= 0:
            return result(
                False,
                "result_unconfirmed",
                "Requested UI condition was not confirmed before the timeout; "
                "inspect the latest observation before deciding whether to repeat an action",
                data=evidence(),
            )
        sleep(min(interval, remaining))


def wait_for_control(
    observe: Probe,
    state: str = "exists",
    *,
    timeout_seconds: float = 10.0,
    interval_seconds: float = 0.2,
    clock: Callable[[], float] = time.monotonic,
    sleep: Callable[[float], None] = time.sleep,
) -> dict:
    """Wait for a uniquely resolved control to exist, be absent, or be enabled.

    ``enabled`` confirms the UIA enabled property, not visibility or clickability.
    Missing controls may appear on a later probe. Ambiguous selectors and invalid
    parent windows are terminal errors and never count as confirmed absence.
    """
    if state not in {"exists", "absent", "enabled"}:
        return _invalid("state must be exists, absent, or enabled")

    def predicate(observation: dict) -> bool:
        if state == "absent":
            return not observation["exists"]
        if not observation["exists"]:
            return False
        if state == "enabled":
            if not isinstance(observation.get("enabled"), bool):
                raise ValueError("UI probe did not provide a boolean enabled property")
            return observation["enabled"]
        return True

    return _poll(
        observe,
        predicate,
        condition={"target": "control", "state": state},
        timeout_seconds=timeout_seconds,
        interval_seconds=interval_seconds,
        clock=clock,
        sleep=sleep,
    )


def wait_for_text(
    observe: Probe,
    expected_text: str | None = None,
    *,
    match: str = "exact",
    baseline_text: str | None = None,
    timeout_seconds: float = 10.0,
    interval_seconds: float = 0.2,
    clock: Callable[[], float] = time.monotonic,
    sleep: Callable[[float], None] = time.sleep,
) -> dict:
    """Verify UI text equality, containment, or change from an explicit baseline.

    Text matching is case-sensitive. The platform adapter decides which UIA
    text property to read and records that choice in its observation. A missing
    control never proves a text change, including when the baseline is empty.
    """
    if match not in {"exact", "contains", "changed"}:
        return _invalid("match must be exact, contains, or changed")
    if match == "changed" and not isinstance(baseline_text, str):
        return _invalid("baseline_text is required for changed text verification")
    if match != "changed" and not isinstance(expected_text, str):
        return _invalid("expected_text is required for exact or contains verification")
    condition = {"target": "text", "match": match}
    if match == "changed":
        condition["baseline_text"] = baseline_text
    else:
        condition["expected_text"] = expected_text

    def predicate(observation: dict) -> bool:
        if not observation["exists"]:
            return False
        value = observation.get("text")
        if value is None:
            return False
        if not isinstance(value, str):
            raise ValueError("UI probe did not provide a text property")
        if match == "changed":
            return value != baseline_text
        if match == "contains":
            return expected_text in value
        return value == expected_text

    return _poll(
        observe,
        predicate,
        condition=condition,
        timeout_seconds=timeout_seconds,
        interval_seconds=interval_seconds,
        clock=clock,
        sleep=sleep,
    )


def wait_for_window(
    observe: Probe,
    state: str = "exists",
    *,
    baseline: dict | None = None,
    timeout_seconds: float = 10.0,
    interval_seconds: float = 0.2,
    clock: Callable[[], float] = time.monotonic,
    sleep: Callable[[float], None] = time.sleep,
) -> dict:
    """Verify window existence, absence, or a change in supplied baseline fields.

    ``baseline`` for changed windows contains at least one of handle, title,
    class_name, process_id, bounds, enabled, or visible. A full observation with
    ``exists=True`` is also accepted. Only supplied fields are compared; absence
    must be checked explicitly with ``state='absent'``.
    """
    if state not in {"exists", "absent", "changed"}:
        return _invalid("state must be exists, absent, or changed")
    expected = None
    if state == "changed":
        if not isinstance(baseline, Mapping):
            return _invalid("baseline is required for changed window verification")
        if "exists" in baseline and baseline["exists"] is not True:
            return _invalid("changed window verification requires an existing baseline window")
        if set(baseline) - WINDOW_FIELDS - {"exists"}:
            return _invalid("baseline contains unsupported window properties")
        expected = deepcopy({key: value for key, value in baseline.items() if key != "exists"})
        if not expected:
            return _invalid("baseline must contain at least one comparable window property")
    condition = {"target": "window", "state": state}
    if expected is not None:
        condition["baseline"] = expected

    def predicate(observation: dict) -> bool:
        if state == "absent":
            return not observation["exists"]
        if not observation["exists"]:
            return False
        if state == "changed":
            if any(key not in observation for key in expected):
                raise ValueError("UI probe did not provide all baseline window properties")
            return any(observation[key] != value for key, value in expected.items())
        return True

    return _poll(
        observe,
        predicate,
        condition=condition,
        timeout_seconds=timeout_seconds,
        interval_seconds=interval_seconds,
        clock=clock,
        sleep=sleep,
    )
