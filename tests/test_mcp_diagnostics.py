"""Tests for full Codex setup diagnostics."""

from __future__ import annotations

import importlib.util
import json
from collections import deque
from pathlib import Path
from types import SimpleNamespace

import pytest

# Diagnostics only needs the standard library. Load its real source directly so
# these tests also run on Linux without importing the package's Windows GUI API.
_spec = importlib.util.spec_from_file_location(
    "standalone_diagnostics",
    Path(__file__).parents[1] / "src" / "easy_uiauto" / "mcp" / "diagnostics.py",
)
diagnostics = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(diagnostics)


def test_python_vision_dependencies_are_not_reinstalled_when_present(monkeypatch) -> None:
    monkeypatch.setattr(diagnostics.importlib.util, "find_spec", lambda _module: object())
    monkeypatch.setattr(
        diagnostics.subprocess,
        "run",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(
            AssertionError("pip was not expected")
        ),
    )

    result = diagnostics.ensure_python_vision_dependencies("1.2.3")

    assert result["ok"] is True
    assert "already installed" in result["detail"]


def test_existing_tesseract_is_added_to_user_path(monkeypatch, tmp_path) -> None:
    executable = tmp_path / "Tesseract-OCR" / "tesseract.exe"
    calls: list[Path] = []
    monkeypatch.setattr(diagnostics, "_find_tesseract", lambda: executable)
    monkeypatch.setattr(diagnostics, "_add_to_user_path", calls.append)

    result, path = diagnostics.ensure_tesseract()

    assert result["ok"] is True
    assert path == str(executable)
    assert calls == [executable.parent]


def test_missing_tesseract_is_installed_with_winget(monkeypatch, tmp_path) -> None:
    executable = tmp_path / "Tesseract-OCR" / "tesseract.exe"
    discoveries = deque([None, executable])
    commands: list[list[str]] = []
    monkeypatch.setattr(diagnostics, "_find_tesseract", discoveries.popleft)
    monkeypatch.setattr(
        diagnostics.shutil,
        "which",
        lambda name: "winget.exe" if name == "winget" else None,
    )
    monkeypatch.setattr(
        diagnostics.subprocess,
        "run",
        lambda command, **_kwargs: commands.append(command)
        or SimpleNamespace(returncode=0, stdout="", stderr=""),
    )
    monkeypatch.setattr(diagnostics, "_add_to_user_path", lambda _path: None)

    result, path = diagnostics.ensure_tesseract()

    assert result["ok"] is True
    assert path == str(executable)
    assert commands == [
        [
            "winget.exe",
            "install",
            "--id",
            diagnostics.TESSERACT_PACKAGE_ID,
            "--exact",
            "--silent",
            "--accept-package-agreements",
            "--accept-source-agreements",
            "--disable-interactivity",
        ]
    ]


def test_report_formats_pass_and_failure() -> None:
    report = diagnostics.format_report(
        [
            {"name": "UIA", "ok": True, "timing_ms": 1.2, "detail": "ready"},
            {"name": "OCR", "ok": False, "timing_ms": 3.4, "detail": "failed"},
        ]
    )

    assert "[PASS] UIA: 1.2 ms - ready" in report
    assert "[FAIL] OCR: 3.4 ms - failed" in report


@pytest.fixture
def available_local_dependencies(monkeypatch):
    """Simulate prerequisites only; this fixture does not verify Windows UIA."""
    monkeypatch.setattr(diagnostics, "_modules_missing", lambda _modules: [])
    monkeypatch.setattr(diagnostics, "_readiness_probe", lambda _name, _timeout: None)
    monkeypatch.setattr(diagnostics, "_ocr_languages", lambda _timeout: ({"eng"}, None))


def test_readiness_reports_unsupported_uia_platform(monkeypatch, available_local_dependencies):
    monkeypatch.setattr(diagnostics.sys, "platform", "linux")

    result = diagnostics.get_capability_readiness(vision_settings={})

    assert result["capabilities"]["uia"]["available"] is False
    assert result["capabilities"]["uia"]["missing"] == ["Windows"]
    assert result["capabilities"]["image_matching"]["available"] is True
    assert result["capabilities"]["image_matching"]["runtime_verified"] is False
    assert result["screen_capture_checked"] is False


def test_remote_configured_is_not_runtime_verified_or_disclosed(available_local_dependencies):
    values = {
        "EASY_UIAUTO_VISION_API_URL": "https://user:secret-url@example.test/chat?key=secret-query",
        "EASY_UIAUTO_VISION_API_KEY": "secret-token",
        "EASY_UIAUTO_VISION_MODEL": "private-model",
    }

    result = diagnostics.get_capability_readiness(vision_settings=values)
    remote = result["capabilities"]["remote_vision"]

    assert remote["configured"] is True
    assert remote["available"] is None
    assert remote["runtime_verified"] is False
    assert remote["status"] == "unverified"
    assert result["network_checked"] is False
    serialized = json.dumps(result)
    for value in values.values():
        assert value not in serialized
    for sensitive in ("secret-url", "secret-query", "secret-token", "private-model"):
        assert sensitive not in serialized


def test_remote_reports_missing_settings_and_invalid_url(available_local_dependencies):
    result = diagnostics.get_capability_readiness(
        vision_settings={"EASY_UIAUTO_VISION_API_URL": "file:///secret"}
    )["capabilities"]["remote_vision"]

    assert result["configured"] is False
    assert result["available"] is False
    assert result["missing"] == ["EASY_UIAUTO_VISION_API_KEY", "EASY_UIAUTO_VISION_MODEL"]
    assert result["issues"] == ["invalid_api_url"]


def test_ocr_reports_missing_requested_language(available_local_dependencies):
    result = diagnostics.get_capability_readiness(
        language="eng+chi_sim", vision_settings={}
    )["capabilities"]["ocr"]

    assert result["available"] is False
    assert result["configured"] is False
    assert result["missing"] == ["language:chi_sim"]
    assert result["installed_languages"] == ["eng"]


def test_missing_modules_and_executable_are_actionable(monkeypatch):
    monkeypatch.setattr(diagnostics.importlib.util, "find_spec", lambda _name: None)
    monkeypatch.setattr(diagnostics.shutil, "which", lambda _name: None)
    monkeypatch.setattr(
        diagnostics.subprocess, "run",
        lambda *_args, **_kwargs: pytest.fail("No executable or imports should be probed"),
    )

    result = diagnostics.get_capability_readiness(vision_settings={})["capabilities"]

    assert result["ocr"]["missing"] == ["pytesseract", "PIL"]
    assert result["ocr"]["issues"] == ["tesseract_executable_missing"]
    assert result["image_matching"]["missing"] == ["cv2", "numpy", "PIL"]
    assert result["ocr"]["runtime_verified"] is False


def test_broken_native_import_does_not_report_available(monkeypatch, available_local_dependencies):
    monkeypatch.setattr(
        diagnostics, "_readiness_probe", lambda _name, _timeout: "local_probe_failed"
    )

    result = diagnostics.get_capability_readiness(vision_settings={})["capabilities"]

    assert result["image_matching"]["configured"] is True
    assert result["image_matching"]["available"] is False
    assert result["image_matching"]["issues"] == ["local_probe_failed"]


def test_local_probe_is_bounded_and_discards_output(monkeypatch):
    def timed_out(command, **kwargs):
        assert kwargs["timeout"] == 0.1
        assert kwargs["stdout"] == diagnostics.subprocess.DEVNULL
        assert kwargs["stderr"] == diagnostics.subprocess.DEVNULL
        raise diagnostics.subprocess.TimeoutExpired(command, 0.1)

    monkeypatch.setattr(diagnostics.subprocess, "run", timed_out)

    assert diagnostics._readiness_probe("image_matching", 0.1) == "probe_timeout"


def test_tesseract_uses_existing_command_and_filters_output(monkeypatch):
    monkeypatch.setitem(
        diagnostics.sys.modules, "pytesseract.pytesseract",
        SimpleNamespace(tesseract_cmd="custom-tesseract"),
    )
    monkeypatch.setattr(diagnostics.shutil, "which", lambda name: name)

    def run(command, **kwargs):
        assert command == ["custom-tesseract", "--list-langs"]
        assert kwargs["timeout"] == 0.2
        return SimpleNamespace(
            returncode=0,
            stdout='List of available languages in "/private/config" (2):\neng\n',
            stderr="script/HanS\n",
        )

    monkeypatch.setattr(diagnostics.subprocess, "run", run)

    assert diagnostics._ocr_languages(0.2) == ({"eng", "script/HanS"}, None)


@pytest.mark.parametrize("timeout", [0, -1, float("nan"), float("inf"), 31])
def test_invalid_probe_timeout_is_rejected(timeout):
    with pytest.raises(ValueError, match="probe_timeout"):
        diagnostics.get_capability_readiness(probe_timeout=timeout)


@pytest.mark.parametrize("language", ["", "--help", "eng chi_sim", "../../secret"])
def test_invalid_language_is_rejected(language):
    with pytest.raises(ValueError, match="language"):
        diagnostics.get_capability_readiness(language=language)
