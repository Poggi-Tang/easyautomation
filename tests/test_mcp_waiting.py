"""Platform-independent tests for bounded, read-only MCP UI condition polling."""

from __future__ import annotations

import importlib
import importlib.util
import sys
from pathlib import Path

import pytest

# Load the optional MCP package without importing the Windows desktop facade.
# A private package name keeps these tests independent of other package tests.
_PACKAGE = "_easy_uiauto_waiting_tests"
_ROOT = Path(__file__).resolve().parents[1] / "src" / "easy_uiauto" / "mcp"
_SPEC = importlib.util.spec_from_file_location(
    _PACKAGE, _ROOT / "__init__.py", submodule_search_locations=[str(_ROOT)]
)
_MODULE = importlib.util.module_from_spec(_SPEC)
sys.modules[_PACKAGE] = _MODULE
_SPEC.loader.exec_module(_MODULE)
waiting = importlib.import_module(f"{_PACKAGE}.waiting")
ActionError = importlib.import_module(f"{_PACKAGE}.results").ActionError


class FakeClock:
    def __init__(self):
        self.value = 0.0
        self.sleeps = []

    def now(self):
        return self.value

    def sleep(self, seconds):
        assert seconds > 0
        self.sleeps.append(seconds)
        self.value += seconds

    def options(self):
        return {"clock": self.now, "sleep": self.sleep}


def samples(*observations):
    iterator = iter(observations)

    def observe():
        item = next(iterator)
        if isinstance(item, Exception):
            raise item
        return item

    return observe


def test_control_wait_observes_missing_then_present_without_repeating_action():
    clock = FakeClock()
    outcome = waiting.wait_for_control(
        samples({"exists": False}, {"exists": True, "enabled": True}),
        **clock.options(),
    )
    assert outcome["ok"] is True
    assert outcome["code"] == "verified"
    assert outcome["verified"] is True
    assert outcome["data"]["attempts"] == 2
    assert outcome["data"]["elapsed_seconds"] == 0.2


def test_typed_control_not_found_is_retried_for_appearance():
    clock = FakeClock()
    outcome = waiting.wait_for_control(
        samples(ActionError("control_not_found", "No match"), {"exists": True}),
        **clock.options(),
    )
    assert outcome["ok"] is True
    assert outcome["data"]["attempts"] == 2


def test_control_disappearance_requires_observed_absence():
    clock = FakeClock()
    outcome = waiting.wait_for_control(
        samples({"exists": True}, {"exists": False}), "absent", **clock.options()
    )
    assert outcome["ok"] is True
    assert outcome["data"]["observation"] == {"exists": False}


@pytest.mark.parametrize("code", ["ambiguous_match", "window_invalid", "action_failed"])
def test_failed_probe_is_terminal_and_does_not_prove_disappearance(code):
    clock = FakeClock()
    outcome = waiting.wait_for_control(
        samples(ActionError(code, "Cannot inspect target")), "absent", **clock.options()
    )
    assert outcome["ok"] is False
    assert outcome["verified"] is False
    assert outcome["code"] == code
    assert outcome["data"]["attempts"] == 1
    assert clock.sleeps == []


def test_unexpected_probe_error_does_not_prove_disappearance():
    outcome = waiting.wait_for_control(
        samples(RuntimeError("COM failed")), "absent", timeout_seconds=0
    )
    assert outcome["code"] == "action_failed"
    assert outcome["verified"] is False


def test_enabled_wait_checks_presence_then_enabled_property():
    clock = FakeClock()
    outcome = waiting.wait_for_control(
        samples(
            {"exists": False},
            {"exists": True, "enabled": False},
            {"exists": True, "enabled": True},
        ),
        "enabled",
        **clock.options(),
    )
    assert outcome["ok"] is True
    assert outcome["data"]["attempts"] == 3


def test_timeout_preserves_last_observation_and_does_not_suggest_repeating_action():
    clock = FakeClock()
    outcome = waiting.wait_for_control(
        lambda: {"exists": False},
        timeout_seconds=0.5,
        interval_seconds=0.3,
        **clock.options(),
    )
    assert outcome["ok"] is False
    assert outcome["code"] == "result_unconfirmed"
    assert outcome["retryable"] is False
    assert outcome["data"]["last_code"] == "control_not_found"
    assert outcome["data"]["observation"] == {"exists": False}
    assert outcome["data"]["attempts"] == 3
    assert clock.sleeps == [0.3, 0.2]
    assert clock.value == 0.5


def test_zero_timeout_performs_one_check_without_sleeping():
    clock = FakeClock()
    outcome = waiting.wait_for_control(
        lambda: {"exists": False}, timeout_seconds=0, **clock.options()
    )
    assert outcome["code"] == "result_unconfirmed"
    assert outcome["data"]["attempts"] == 1
    assert clock.sleeps == []


def test_slow_probe_does_not_claim_success_before_deadline():
    clock = FakeClock()

    def slow_probe():
        clock.value += 2
        return {"exists": True}

    outcome = waiting.wait_for_control(slow_probe, timeout_seconds=1, **clock.options())
    assert outcome["code"] == "result_unconfirmed"
    assert outcome["data"]["elapsed_seconds"] == 2
    assert clock.sleeps == []


@pytest.mark.parametrize(
    "options",
    [
        {"timeout_seconds": -1},
        {"timeout_seconds": 121},
        {"timeout_seconds": float("nan")},
        {"timeout_seconds": float("inf")},
        {"interval_seconds": 0},
        {"interval_seconds": -1},
        {"interval_seconds": float("inf")},
        {"interval_seconds": "bad"},
    ],
)
def test_invalid_polling_options_do_not_touch_desktop(options):
    def unexpected_probe():
        raise AssertionError("Invalid parameters must be rejected before probing")

    outcome = waiting.wait_for_control(unexpected_probe, **options)
    assert outcome["code"] == "invalid_argument"


@pytest.mark.parametrize("observation", [None, {}, {"exists": 1}, {"exists": "false"}])
def test_malformed_probe_output_is_never_absence(observation):
    outcome = waiting.wait_for_control(lambda: observation, "absent", timeout_seconds=0)
    assert outcome["code"] == "action_failed"


def test_missing_enabled_property_does_not_assume_enabled():
    outcome = waiting.wait_for_control(
        lambda: {"exists": True}, "enabled", timeout_seconds=0
    )
    assert outcome["code"] == "action_failed"


@pytest.mark.parametrize(
    ("match", "value", "expected"),
    [("exact", "Saved", "Saved"), ("contains", "Document Saved", "Saved"), ("exact", "", "")],
)
def test_text_verification_observes_expected_content(match, value, expected):
    outcome = waiting.wait_for_text(
        lambda: {"exists": True, "text": value},
        expected,
        match=match,
        timeout_seconds=0,
    )
    assert outcome["ok"] is True
    assert outcome["verified"] is True


def test_text_wait_retries_until_text_changes_from_explicit_baseline():
    clock = FakeClock()
    outcome = waiting.wait_for_text(
        samples(
            {"exists": False},
            {"exists": True, "text": "Loading"},
            {"exists": True, "text": "Ready"},
        ),
        match="changed",
        baseline_text="Loading",
        **clock.options(),
    )
    assert outcome["ok"] is True
    assert outcome["data"]["attempts"] == 3
    assert outcome["data"]["observation"]["text"] == "Ready"


def test_disappearance_never_confirms_text_change():
    outcome = waiting.wait_for_text(
        lambda: {"exists": False}, match="changed", baseline_text="", timeout_seconds=0
    )
    assert outcome["code"] == "result_unconfirmed"


@pytest.mark.parametrize("observation", [{"exists": True}, {"exists": True, "text": None}])
def test_unreadable_text_is_unconfirmed_even_for_empty_expected_text(observation):
    outcome = waiting.wait_for_text(lambda: observation, "", timeout_seconds=0)
    assert outcome["code"] == "result_unconfirmed"
    assert outcome["verified"] is False
    assert outcome["retryable"] is False


def test_initially_unreadable_text_can_be_verified_after_later_read():
    clock = FakeClock()
    outcome = waiting.wait_for_text(
        samples({"exists": True}, {"exists": True, "text": "Ready"}),
        "Ready",
        **clock.options(),
    )
    assert outcome["code"] == "verified"
    assert outcome["data"]["attempts"] == 2


def test_window_change_compares_only_requested_baseline_properties():
    clock = FakeClock()
    outcome = waiting.wait_for_window(
        samples(
            {"exists": True, "title": "Editor", "bounds": {"left": 10}},
            {"exists": True, "title": "Saved - Editor", "bounds": {"left": 10}},
        ),
        "changed",
        baseline={"exists": True, "title": "Editor"},
        **clock.options(),
    )
    assert outcome["ok"] is True
    assert outcome["data"]["attempts"] == 2


def test_window_change_requires_observation_of_baseline_fields():
    outcome = waiting.wait_for_window(
        lambda: {"exists": True, "title": "Editor"},
        "changed",
        baseline={"handle": 17},
        timeout_seconds=0,
    )
    assert outcome["code"] == "action_failed"


def test_missing_window_never_confirms_window_property_change():
    outcome = waiting.wait_for_window(
        lambda: {"exists": False},
        "changed",
        baseline={"title": "Editor"},
        timeout_seconds=0,
    )
    assert outcome["code"] == "result_unconfirmed"


def test_window_disappearance_can_be_verified_explicitly():
    outcome = waiting.wait_for_window(
        lambda: {"exists": False}, "absent", timeout_seconds=0
    )
    assert outcome["ok"] is True
    assert outcome["verified"] is True


@pytest.mark.parametrize("baseline", [None, {}, {"exists": False}, {"exists": True}, {"typo": 1}])
def test_window_change_rejects_missing_or_unknown_baseline(baseline):
    outcome = waiting.wait_for_window(
        lambda: {"exists": True}, "changed", baseline=baseline, timeout_seconds=0
    )
    assert outcome["code"] == "invalid_argument"


def test_text_change_requires_explicit_baseline():
    outcome = waiting.wait_for_text(
        lambda: {"exists": True, "text": "Ready"}, match="changed", timeout_seconds=0
    )
    assert outcome["code"] == "invalid_argument"


def test_empty_text_baseline_is_valid():
    outcome = waiting.wait_for_text(
        lambda: {"exists": True, "text": "Ready"},
        match="changed",
        baseline_text="",
        timeout_seconds=0,
    )
    assert outcome["ok"] is True


def test_mutated_observation_does_not_change_returned_evidence():
    observation = {"exists": True, "bounds": {"left": 1}}
    outcome = waiting.wait_for_window(lambda: observation, timeout_seconds=0)
    observation["bounds"]["left"] = 2
    assert outcome["data"]["observation"]["bounds"]["left"] == 1
